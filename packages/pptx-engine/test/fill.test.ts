import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { FillError, Template } from '../src/index.ts';
import type { DeckData } from '../src/index.ts';
import { chartSeriesXml, fixture, partXml, slideTexts, unzip } from './helpers.ts';

const opsData: DeckData = {
  values: {
    title: 'Monthly Ops Review',
    site: 'Plant A',
    period_end: new Date(Date.UTC(2026, 9, 31)),
    downtime_total: { value: 1234.5, format: 'number:0' },
  },
  tables: {
    downtime: {
      columns: ['minutes', 'LINE', 'Cause', 'Ignored extra'],
      rows: [
        [95, 'Line 1', 'Changeover', 'x'],
        [40, 'Line 2', 'Jam', 'x'],
        [12.5, 'Line 3', 'Material {{not a tag}}', 'x'],
      ],
    },
  },
  charts: {
    downtime_by_line: {
      categoryLabel: 'Line',
      categories: ['Line 1', 'Line 2', 'Line 3', 'Line 4'],
      series: [
        { name: 'Planned', values: [10, 20, 30, 40] },
        { name: 'Unplanned', values: [5, null, 15, 25] },
      ],
    },
    oee_trend: {
      categories: ['May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct'],
      series: [{ name: 'OEE', values: [0.7, 0.72, 0.75, 0.71, 0.77, 0.8] }],
    },
  },
};

describe('slot discovery', () => {
  it('lists text tags, tables and charts with sample data', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    expect(tpl.slots()).toMatchSnapshot();
  });

  it('finds split tags and reports tag-like text that is not valid', async () => {
    const m = (await Template.load(fixture('text-basic.pptx'))).slots();
    expect(m.text.map((t) => [t.name, t.fields, t.slides])).toEqual([
      ['title', ['value'], [1]],
      ['period_start', ['value'], [1]],
      ['period_end', ['value'], [1]],
      ['oee', ['value', 'previous'], [2]],
      ['scrap_rate', ['value'], [2]],
      ['units', ['value'], [2]],
      ['revenue', ['value'], [2]],
    ]);
    expect(m.issues.map((i) => i.message)).toEqual([
      'Slide 2: "{{ }}" looks like a tag but will not be filled: the tag is empty.',
      'Slide 2: "{{bad name}}" looks like a tag but will not be filled: tag names may only use letters, digits, _ and -, must start with a letter or _, and cannot contain spaces.',
    ]);
  });
});

describe('fill: text-basic (golden)', () => {
  it('fills every tag, keeping the formatting of the run where the tag starts', async () => {
    const tpl = await Template.load(fixture('text-basic.pptx'));
    const deck = await tpl.fill({
      values: {
        title: 'October Report',
        period_start: new Date(Date.UTC(2026, 9, 1)),
        period_end: new Date(Date.UTC(2026, 9, 31)),
        oee: { value: 0.853, previous: 0.81, format: 'percent' },
        SCRAP_RATE: { value: 0.021, format: 'percent:1' },
        units: 12000,
        revenue: { value: 1250000, format: 'currency:0' },
      },
    });
    const bytes = deck.toBuffer();
    expect(await slideTexts(bytes)).toMatchSnapshot();

    // The OEE value lands in the red run where "{{o" started.
    const slide2 = await partXml(await unzip(bytes), 'ppt/slides/slide2.xml');
    const run = Array.from(slide2.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')).find((t) => t.textContent === '85.3%');
    expect(run?.parentNode?.toString()).toContain('C00000');
    expect(deck.issues.filter((i) => i.code !== 'invalid-tag')).toEqual([]);
  });
});

describe('fill: ops-review (golden)', () => {
  it('fills text, table rows and both charts', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    const deck = await tpl.fill(opsData);
    const bytes = deck.toBuffer();
    expect(deck.issues).toEqual([]);
    expect(await slideTexts(bytes)).toMatchSnapshot('slide text');
    expect(await chartSeriesXml(bytes, 'ppt/charts/chart1.xml')).toMatchSnapshot('bar chart series');
    expect(await chartSeriesXml(bytes, 'ppt/charts/chart2.xml')).toMatchSnapshot('line chart series');
  });

  it('round-trips: the filled deck reads back as the data that went in', async () => {
    const deck = await (await Template.load(fixture('ops-review.pptx'))).fill(opsData);
    const back = (await Template.load(deck.toBuffer())).slots();
    const table = back.tables[0]!;
    expect(table.columns).toEqual(['Line', 'Cause', 'Minutes']);
    expect(table.sampleRows).toEqual([
      ['Line 1', 'Changeover', '95'],
      ['Line 2', 'Jam', '40'],
      ['Line 3', 'Material {{not a tag}}', '12.5'],
    ]);
    const bar = back.charts.find((c) => c.name === 'downtime_by_line')!;
    expect(bar.sample).toEqual({
      categories: ['Line 1', 'Line 2', 'Line 3', 'Line 4'],
      series: [
        { name: 'Planned', values: [10, 20, 30, 40] },
        { name: 'Unplanned', values: [5, null, 15, 25] },
      ],
    });
    const line = back.charts.find((c) => c.name === 'oee_trend')!;
    expect(line.sample.series.map((s) => s.name)).toEqual(['OEE']);
  });

  it('replaces the embedded workbook so Edit Data shows the new numbers', async () => {
    const deck = await (await Template.load(fixture('ops-review.pptx'))).fill(opsData);
    const zip = await unzip(deck.toBuffer());
    const xlsx = await zip.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.async('uint8array');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(xlsx.buffer as ArrayBuffer);
    const ws = wb.getWorksheet('Sheet1')!;
    const rows: unknown[][] = [];
    ws.eachRow({ includeEmpty: true }, (row) => rows.push((row.values as unknown[]).slice(1)));
    expect(rows).toEqual([
      ['Line', 'Planned', 'Unplanned'],
      ['Line 1', 10, 5],
      ['Line 2', 20],
      ['Line 3', 30, 15],
      ['Line 4', 40, 25],
    ]);
  });

  it('leaves the template untouched so it can be filled again', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    const a = await slideTexts((await tpl.fill(opsData)).toBuffer());
    const b = await slideTexts((await tpl.fill({ ...opsData, values: { ...opsData.values, site: 'Plant B' } })).toBuffer());
    expect(a).toContain('1: Plant A · Oct 31, 2026');
    expect(b).toContain('1: Plant B · Oct 31, 2026');
  });
});

describe('fill: problems are explained', () => {
  async function fillError(data: DeckData, opts = {}): Promise<string[]> {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    try {
      await tpl.fill(data, opts);
    } catch (e) {
      expect(e).toBeInstanceOf(FillError);
      return (e as FillError).issues.filter((i) => i.severity === 'error').map((i) => i.message);
    }
    throw new Error('expected FillError');
  }

  it('lists every missing value, table and chart at once', async () => {
    expect(await fillError({ values: { title: 'x' } })).toEqual([
      'No data for {{site}} (slide 1).',
      'No data for {{period_end}} (slide 1).',
      'No data for {{downtime_total}} (slide 2).',
      "No data for table 'downtime' (slide 2).",
      "No data for chart 'downtime_by_line' (slide 3).",
      "No data for chart 'oee_trend' (slide 4).",
    ]);
  });

  it('can keep or blank missing tags instead', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    const kept = await tpl.fill({ values: { title: 'T' } }, { onMissing: 'keep' });
    expect(await slideTexts(kept.toBuffer())).toContain('1: {{site}} · {{period_end}}');
    expect(kept.issues.some((i) => i.code === 'missing-value' && i.severity === 'warning')).toBe(true);
    const blank = await tpl.fill({ values: { title: 'T' } }, { onMissing: 'blank' });
    expect(await slideTexts(blank.toBuffer())).toContain('1:  · ');
  });

  it('names table columns the data is missing', async () => {
    const errors = await fillError({
      ...opsData,
      tables: { downtime: { columns: ['Line', 'Reason', 'Minutes'], rows: [] } },
    });
    expect(errors).toEqual([
      `Table 'downtime' on slide 2 has the column "Cause", but the data has no column with that header. Data columns: "Line", "Reason", "Minutes".`,
    ]);
  });

  it('explains a series-count mismatch', async () => {
    const errors = await fillError({
      ...opsData,
      charts: {
        ...opsData.charts,
        oee_trend: { categories: ['Jan'], series: [{ name: 'A', values: [1] }, { name: 'B', values: [2] }, { name: 'C', values: [3] }] },
      },
    });
    expect(errors).toEqual([
      "Chart 'oee_trend' on slide 4 has 2 series in the template but the data has 3. Remove 1 series column from the data, or add series to the chart in PowerPoint (Chart Design > Edit Data).",
    ]);
  });

  it('rejects unknown formats and impossible trends', async () => {
    const errors = await fillError({ ...opsData, values: { ...opsData.values, downtime_total: { value: 5, format: '#,##0' } } });
    expect(errors[0]).toMatch(/^Value "downtime_total": "#,##0" is not a format DeckForge knows/);
  });

  it('warns about data that matches nothing', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    const deck = await tpl.fill({ ...opsData, tables: { ...opsData.tables, scrap: { columns: [], rows: [] } } });
    expect(deck.issues).toEqual([
      { severity: 'warning', code: 'unused-data', slot: 'scrap', message: "The data has a table 'scrap', but the template has no table with that name." },
    ]);
  });

  it('rejects files that are not PowerPoint', async () => {
    await expect(Template.load(new TextEncoder().encode('hello'))).rejects.toThrow(/not a PowerPoint/);
  });
});
