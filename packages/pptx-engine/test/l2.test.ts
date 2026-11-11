import { describe, expect, it } from 'vitest';
import { FillError, Template } from '../src/index.ts';
import type { DeckData } from '../src/index.ts';
import { fixture, partXml, slideTexts, unzip } from './helpers.ts';

const base: DeckData = {
  values: {
    title: 'Q3 Review',
    presenter: 'Dana',
    company: 'Acme',
    incidents: 2,
    skip_appendix: 'no',
  },
  rows: {
    sites: [
      { site: 'Plant A', oee: { value: 0.8, previous: 0.7, format: 'percent' }, manager: 'Ann' },
      { site: 'Plant B', oee: { value: 0.6, previous: 0.7, format: 'percent' }, manager: 'Bo' },
      { site: 'Plant C', oee: { value: 0.7, previous: 0.7, format: 'percent' }, manager: 'Cy' },
    ],
  },
  charts: {
    site_downtime: { categories: ['Planned', 'Unplanned', 'Other'], series: [{ name: 'Hours', values: [3, 5, 1] }] },
    scrap_share: { categories: ['Setup', 'Material'], series: [{ name: 'Scrap', values: [4, 6] }] },
    volume_yield: {
      categories: ['Jul', 'Aug', 'Sep', 'Oct'],
      series: [
        { name: 'Volume', values: [1, 2, 3, 4] },
        { name: 'Yield', values: [0.9, 0.91, 0.92, 0.93] },
      ],
    },
  },
};

async function fill(data: DeckData = base): Promise<Uint8Array> {
  const tpl = await Template.load(fixture('l2-features.pptx'));
  return (await tpl.fill(data)).toBuffer();
}

async function partNames(bytes: Uint8Array, re: RegExp): Promise<string[]> {
  return Object.keys((await unzip(bytes)).files).filter((p) => re.test(p)).sort();
}

describe('manifest: L2 slots', () => {
  it('reports repeats, rules, layout tags and chart kinds', async () => {
    const m = (await Template.load(fixture('l2-features.pptx'))).slots();
    expect(m.repeats).toEqual([
      { kind: 'repeat', name: 'sites', slides: [2], tags: [{ name: 'site', fields: ['value'] }, { name: 'oee', fields: ['value', 'trend'] }, { name: 'manager', fields: ['value'] }] },
    ]);
    expect(m.rules).toEqual([
      { slide: 2, kind: 'repeat', target: 'sites' },
      { slide: 4, kind: 'show-if', target: 'incidents' },
      { slide: 5, kind: 'hide-if', target: 'skip_appendix' },
    ]);
    expect(m.text.find((t) => t.name === 'company')).toMatchObject({ slides: [], layouts: ['BRANDED'] });
    expect(m.charts.map((c) => [c.name, c.chartType])).toEqual([
      ['site_downtime', 'pie'],
      ['scrap_share', 'doughnut'],
      ['volume_yield', 'combo'],
    ]);
  });
});

describe('fill: repeat slides', () => {
  it('makes one copy per row, filled from the row, with its own chart', async () => {
    const bytes = await fill();
    const texts = await slideTexts(bytes);
    expect(texts.filter((t) => /^\d: Plant/.test(t))).toEqual(['2: Plant A', '3: Plant B', '4: Plant C']);
    expect(texts).toContain('3: OEE 60% (▼) · Manager: Bo');
    expect(texts).toContain('4: OEE 70% (▶) · Manager: Cy');

    // 7 slides: title, 3 sites, charts, safety, appendix. Each site copy has its own chart part + workbook.
    expect(await partNames(bytes, /^ppt\/slides\/slide\d+\.xml$/)).toHaveLength(7);
    expect(await partNames(bytes, /^ppt\/charts\/chart\d+\.xml$/)).toHaveLength(5);
    expect(await partNames(bytes, /^ppt\/embeddings\/.*\.xlsx$/)).toHaveLength(5);

    // Chart titles get the row's tags.
    const back = (await Template.load(bytes)).slots();
    expect(back.charts.filter((c) => c.name === 'site_downtime')).toHaveLength(1); // copies share the name
    const zip = await unzip(bytes);
    const titles: string[] = [];
    for (const p of await partNames(bytes, /^ppt\/charts\/chart\d+\.xml$/)) {
      const t = (await partXml(zip, p)).getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'title')[0];
      if (t) titles.push(Array.from(t.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 't')).map((x) => x.textContent).join(''));
    }
    expect(titles.sort()).toEqual(['Plant A downtime', 'Plant B downtime', 'Plant C downtime']);
  });

  it('keeps notes text, fills notes tags and removes directive lines', async () => {
    const zip = await unzip(await fill());
    const notes: string[] = [];
    for (const p of Object.keys(zip.files).filter((f) => /notesSlide\d+\.xml$/.test(f))) {
      notes.push((await partXml(zip, p)).documentElement!.textContent ?? '');
    }
    const all = notes.join('|');
    expect(all).toContain('Presenter: Dana');
    expect(all).toContain('One slide per site.');
    expect(all).not.toContain('#repeat');
    expect(all).not.toContain('#show-if');
    expect(all).toContain('#showif typo'); // not a directive: left alone
  });

  it('removes the slide (and its parts) for zero rows', async () => {
    const bytes = await fill({ ...base, rows: { sites: [] } });
    expect((await slideTexts(bytes)).some((t) => t.includes('Plant'))).toBe(false);
    expect(await partNames(bytes, /^ppt\/slides\/slide\d+\.xml$/)).toHaveLength(4);
    expect(await partNames(bytes, /^ppt\/charts\/chart\d+\.xml$/)).toHaveLength(2);
    expect(await partNames(bytes, /^ppt\/notesSlides\/.*\.xml$/)).toHaveLength(4);
    const ct = (await (await unzip(bytes)).file('[Content_Types].xml')!.async('string'));
    for (const m of ct.matchAll(/PartName="\/([^"]+)"/g)) {
      expect((await unzip(bytes)).file(m[1]!), `override for missing part ${m[1]}`).not.toBeNull();
    }
  });

  it('reports missing rows with the template slide number', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const err = await tpl.fill({ ...base, rows: {} }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(FillError);
    expect((err as FillError).issues.filter((i) => i.severity === 'error').map((i) => i.message)).toEqual([
      'No rows for "#repeat sites" (slide 2).',
      'No data for {{site}} (slide 2).',
      'No data for {{oee}} (slide 2).',
      'No data for {{manager}} (slide 2).',
    ]);
  });
});

describe('fill: show-if / hide-if', () => {
  it('keeps or drops slides by value', async () => {
    const shown = await slideTexts(await fill());
    expect(shown).toContain('6: Safety incidents: 2');
    expect(shown).toContain('7: Appendix');

    const hidden = await slideTexts(await fill({ ...base, values: { ...base.values, incidents: 0, skip_appendix: 'yes' } }));
    expect(hidden.some((t) => t.includes('Safety') || t.includes('Appendix'))).toBe(false);
  });

  it('drops a show-if slide with a warning when the key is missing', async () => {
    const values = { ...base.values };
    delete values.incidents;
    const deck = await (await Template.load(fixture('l2-features.pptx'))).fill({ ...base, values });
    expect(deck.issues.map((i) => i.message)).toContain('Slide 4: "#show-if incidents" has no data for incidents, so the slide is left out.');
  });
});

describe('fill: layouts, pie, doughnut, combo', () => {
  it('fills tags in non-placeholder shapes on layouts', async () => {
    const zip = await unzip(await fill());
    const layouts = Object.keys(zip.files).filter((p) => /slideLayouts\/slideLayout\d+\.xml$/.test(p));
    const texts = await Promise.all(layouts.map(async (p) => (await partXml(zip, p)).documentElement!.textContent ?? ''));
    expect(texts.join('|')).toContain('Acme confidential');
  });

  it('fills pie/doughnut categories and both groups of a combo', async () => {
    const back = (await Template.load(await fill())).slots();
    expect(back.charts.find((c) => c.name === 'scrap_share')!.sample).toEqual(base.charts!.scrap_share);
    expect(back.charts.find((c) => c.name === 'volume_yield')!.sample).toEqual(base.charts!.volume_yield);
  });

  it('needs enough series to keep each part of a combo', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const err = await tpl
      .fill({ ...base, charts: { ...base.charts, volume_yield: { categories: ['a'], series: [{ name: 'Volume', values: [1] }] } } })
      .catch((e: unknown) => e);
    expect((err as FillError).issues.map((i) => i.message)).toContain(
      "Chart 'volume_yield' on slide 3 combines chart types and needs at least 2 series so each part keeps one; the data has 1.",
    );
  });
});

describe('performance', () => {
  it('fills a 30-slide deck with 10 tables and 20 charts in under 5 seconds', async () => {
    const tpl = await Template.load(fixture('perf-30.pptx'));
    const m = tpl.slots();
    const data: DeckData = {
      values: Object.fromEntries([['title', 'Perf'], ...m.text.filter((t) => t.name !== 'title').map((t) => [t.name, { value: 2, previous: 1 }])]),
      tables: Object.fromEntries(m.tables.map((t) => [t.name, { columns: t.columns, rows: Array.from({ length: 8 }, (_, i) => [`Item ${i}`, i]) }])),
      charts: Object.fromEntries(m.charts.map((c) => [c.name, { categories: ['a', 'b', 'c', 'd', 'e', 'f'], series: [{ name: 'S', values: [1, 2, 3, 4, 5, 6] }] }])),
    };
    const start = performance.now();
    const deck = await tpl.fill(data);
    const ms = performance.now() - start;
    expect(deck.issues).toEqual([]);
    expect(m.slideCount).toBe(30);
    expect(ms).toBeLessThan(5000);
  });
});

describe('trial watermark', () => {
  it('adds a small text box to every slide that is never read as a slot', async () => {
    const tpl = await Template.load(fixture('l2-features.pptx'));
    const bytes = (await tpl.fill(base, { watermark: 'Made with DeckForge' })).toBuffer();
    const texts = await slideTexts(bytes);
    expect(texts.filter((t) => t.endsWith('Made with DeckForge'))).toHaveLength(7);
    const back = (await Template.load(bytes)).slots();
    expect(back.issues.some((i) => /watermark/i.test(i.message))).toBe(false);
    expect([...back.tables, ...back.charts].some((x) => /watermark/i.test(x.name))).toBe(false);
  });
});
