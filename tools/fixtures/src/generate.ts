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

const fixtures: Record<string, () => Pres> = {
  'text-basic': textBasic,
  'ops-review': opsReview,
};

mkdirSync(outDir, { recursive: true });
for (const [name, build] of Object.entries(fixtures)) {
  const buf = (await build().write({ outputType: 'nodebuffer' })) as Uint8Array;
  const file = join(outDir, `${name}.pptx`);
  writeFileSync(file, buf);
  console.log(`wrote ${file} (${buf.byteLength} bytes)`);
}
