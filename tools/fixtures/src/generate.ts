/**
 * Generates synthetic fixture templates with pptxgenjs.
 *
 * These unblock engine tests. They are NOT a substitute for templates authored in real
 * PowerPoint (which add quirks such as spell-check run splits and placeholder inheritance);
 * those live alongside these in fixtures/templates/ with a `ppt-` prefix.
 *
 * Run: pnpm fixtures
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import pptxgenModule from 'pptxgenjs';

// pptxgenjs ships CommonJS-style types for its ESM entry: under nodenext TypeScript sees the
// module namespace while Node loads the class itself. Bridge the two here.
const PptxGenJS = pptxgenModule as unknown as typeof pptxgenModule.default;

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(here, '../../../fixtures/templates');

type Pres = InstanceType<typeof pptxgenModule.default>;

function newPres(title: string): Pres {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title = title;
  pres.author = 'DeckForge fixtures';
  return pres;
}

/** Text tags only, including a tag deliberately split across differently formatted runs. */
function textBasic(): Pres {
  const pres = newPres('text-basic');

  const s1 = pres.addSlide();
  s1.addText('{{title}}', { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 36, bold: true, objectName: 'Title 1' });
  s1.addText('{{period_start}} – {{period_end}}', { x: 0.5, y: 2.6, w: 9, h: 0.6, fontSize: 18, objectName: 'Subtitle 2' });

  const s2 = pres.addSlide();
  s2.addText(
    [
      { text: 'OEE: ', options: { bold: true } },
      { text: '{{o', options: { color: 'C00000' } },
      { text: 'ee}}', options: { color: '0070C0' } },
      { text: ' (was {{oee.previous}})', options: { breakLine: true } },
      { text: 'Scrap rate: {{ scrap_rate }}', options: { breakLine: true } },
      { text: 'Units: {{units}} · Revenue: {{revenue}}' },
    ],
    { x: 0.5, y: 0.5, w: 9, h: 2, fontSize: 20, objectName: 'KPI Box' },
  );
  s2.addText('Not a tag: {single} and {{ }} and {{bad name}}', { x: 0.5, y: 3, w: 9, h: 0.5, fontSize: 12, objectName: 'Footnote' });
  return pres;
}

/** Ops review: text tags, a named table, a bar chart and a line chart. */
function opsReview(): Pres {
  const pres = newPres('ops-review');

  const s1 = pres.addSlide();
  s1.addText('{{title}}', { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 36, bold: true });
  s1.addText('{{site}} · {{period_end}}', { x: 0.5, y: 2.6, w: 9, h: 0.6, fontSize: 18 });

  const s2 = pres.addSlide();
  s2.addText('Downtime by cause', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true });
  const header = { bold: true, color: 'FFFFFF', fill: { color: '1F4E79' } };
  s2.addTable(
    [
      [
        { text: 'Line', options: header },
        { text: 'Cause', options: header },
        { text: 'Minutes', options: { ...header, align: 'right' } },
      ],
      [{ text: 'Line 1' }, { text: 'Changeover' }, { text: '120', options: { align: 'right' } }],
      [{ text: 'Line 2' }, { text: 'Jam' }, { text: '45', options: { align: 'right' } }],
    ],
    { x: 0.5, y: 1.1, w: 9, colW: [2, 5, 2], rowH: 0.4, fontSize: 14, border: { type: 'solid', pt: 1, color: 'BFBFBF' }, objectName: 'downtime' },
  );
  s2.addText('Total: {{downtime_total}} min', { x: 0.5, y: 4.8, w: 9, h: 0.4, fontSize: 14 });

  const s3 = pres.addSlide();
  s3.addText('Downtime by line', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true });
  s3.addChart(
    pres.ChartType.bar,
    [
      { name: 'Planned', labels: ['Line 1', 'Line 2', 'Line 3'], values: [60, 30, 45] },
      { name: 'Unplanned', labels: ['Line 1', 'Line 2', 'Line 3'], values: [120, 45, 20] },
    ],
    { x: 0.5, y: 1.1, w: 9, h: 4, barDir: 'col', barGrouping: 'stacked', showLegend: true, legendPos: 'b', objectName: 'downtime_by_line' },
  );

  const s4 = pres.addSlide();
  s4.addText('OEE trend', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true });
  s4.addChart(
    pres.ChartType.line,
    [
      { name: 'OEE', labels: ['Jan', 'Feb', 'Mar', 'Apr'], values: [0.71, 0.74, 0.69, 0.78] },
      { name: 'Target', labels: ['Jan', 'Feb', 'Mar', 'Apr'], values: [0.75, 0.75, 0.75, 0.75] },
    ],
    { x: 0.5, y: 1.1, w: 9, h: 4, valAxisLabelFormatCode: '0%', dataLabelFormatCode: '0%', showLegend: true, legendPos: 'b', objectName: 'oee_trend' },
  );
  return pres;
}

/** L2 features: pie/doughnut/combo charts, #repeat, #show-if/#hide-if, notes and master tags. */
function l2Features(): Pres {
  const pres = newPres('l2-features');
  pres.defineSlideMaster({
    title: 'BRANDED',
    objects: [
      { text: { text: '{{company}} confidential', options: { x: 0.3, y: 5.2, w: 5, h: 0.3, fontSize: 9, color: '777777' } } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 1.2, w: 9, h: 3.5 }, text: '{{never_shown}}' } },
    ],
  });

  const s1 = pres.addSlide({ masterName: 'BRANDED' });
  s1.addText('{{title}}', { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 30, bold: true });
  s1.addNotes('Presenter: {{presenter}}');

  // Repeated once per row of the "sites" sheet; tags come from the row.
  const s2 = pres.addSlide({ masterName: 'BRANDED' });
  s2.addText('{{site}}', { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 28, bold: true });
  s2.addText('OEE {{oee}} ({{oee.trend}}) · Manager: {{manager}}', { x: 0.5, y: 1.2, w: 9, h: 0.5, fontSize: 18 });
  s2.addChart(
    pres.ChartType.pie,
    [{ name: 'Downtime', labels: ['Planned', 'Unplanned'], values: [30, 70] }],
    { x: 0.5, y: 1.9, w: 4.5, h: 3, showLegend: true, legendPos: 'b', objectName: 'site_downtime', showTitle: true, title: '{{site}} downtime' },
  );
  s2.addNotes('#repeat sites\nOne slide per site.');

  const s3 = pres.addSlide({ masterName: 'BRANDED' });
  s3.addText('Scrap by cause', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true });
  s3.addChart(
    pres.ChartType.doughnut,
    [{ name: 'Scrap', labels: ['Setup', 'Material', 'Other'], values: [5, 3, 2] }],
    { x: 0.5, y: 1.1, w: 4.4, h: 3.8, showLegend: true, legendPos: 'b', objectName: 'scrap_share' },
  );
  // Combo charts: pptxgenjs takes (types, options) at runtime, unlike its typings.
  (s3.addChart as unknown as (types: unknown, opts: unknown) => void)(
    [
      { type: pres.ChartType.bar, data: [{ name: 'Volume', labels: ['Jul', 'Aug', 'Sep'], values: [120, 134, 128] }], options: { barDir: 'col' } },
      { type: pres.ChartType.line, data: [{ name: 'Yield', labels: ['Jul', 'Aug', 'Sep'], values: [0.96, 0.95, 0.97] }], options: { secondaryValAxis: true, secondaryCatAxis: true } },
    ],
    { x: 5.1, y: 1.1, w: 4.4, h: 3.8, showLegend: true, legendPos: 'b', objectName: 'volume_yield', valAxes: [{ showValAxisTitle: false }, { showValAxisTitle: false, valGridLine: { style: 'none' } }], catAxes: [{ catAxisTitle: '' }, { catAxisHidden: true }] },
  );

  const s4 = pres.addSlide({ masterName: 'BRANDED' });
  s4.addText('Safety incidents: {{incidents}}', { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 24, bold: true });
  s4.addNotes('#show-if incidents');

  const s5 = pres.addSlide({ masterName: 'BRANDED' });
  s5.addText('Appendix', { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 24, bold: true });
  s5.addNotes('#hide-if skip_appendix\n#showif typo');
  return pres;
}

/** 30 slides of text, tables and charts, for the "30-slide deck in under 5 seconds" target. */
function perf30(): Pres {
  const pres = newPres('perf-30');
  for (let i = 1; i <= 30; i++) {
    const s = pres.addSlide();
    s.addText(`Slide ${i}: {{title}} · {{kpi_${i}}} ({{kpi_${i}.trend}})`, { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 18 });
    if (i % 3 === 1) {
      s.addTable(
        [['Item', 'Value'].map((t) => ({ text: t })), ...[1, 2, 3].map((r) => [{ text: `Row ${r}` }, { text: String(r) }])],
        { x: 0.5, y: 0.9, w: 9, colW: [6, 3], rowH: 0.3, fontSize: 10, objectName: `table_${i}` },
      );
    } else {
      s.addChart(
        i % 3 === 2 ? pres.ChartType.bar : pres.ChartType.line,
        [{ name: 'Series', labels: ['A', 'B', 'C'], values: [1, 2, 3] }],
        { x: 0.5, y: 0.9, w: 9, h: 4.3, objectName: `chart_${i}` },
      );
    }
  }
  return pres;
}

const fixtures: Record<string, () => Pres> = {
  'text-basic': textBasic,
  'ops-review': opsReview,
  'l2-features': l2Features,
  'perf-30': perf30,
};

mkdirSync(outDir, { recursive: true });
for (const [name, build] of Object.entries(fixtures)) {
  const buf = (await build().write({ outputType: 'nodebuffer' })) as Uint8Array;
  const file = join(outDir, `${name}.pptx`);
  writeFileSync(file, buf);
  console.log(`wrote ${file} (${buf.byteLength} bytes)`);
}
