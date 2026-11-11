/** Reading and filling DrawingML charts (`c:chartSpace`). See design-log 0003, "Chart fill rules". */
import { NS, child, children, create, descendants, path, removeChildren, removeNode, setText, textOf } from './xml.ts';
import type { Document, Element } from './xml.ts';
import type { ChartData, ChartType } from './types.ts';

/** Chart groups DeckForge can fill; all share the category/value series structure. */
const SUPPORTED: Record<string, Exclude<ChartType, 'combo'>> = {
  barChart: 'bar',
  bar3DChart: 'bar',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  doughnutChart: 'doughnut',
};

const TYPE_LABELS: Record<string, string> = {
  bar3DChart: '3-D bar',
  line3DChart: '3-D line',
  pieChart: 'pie',
  pie3DChart: '3-D pie',
  ofPieChart: 'pie-of-pie',
  doughnutChart: 'doughnut',
  areaChart: 'area',
  area3DChart: '3-D area',
  scatterChart: 'scatter',
  bubbleChart: 'bubble',
  radarChart: 'radar',
  stockChart: 'stock',
  surfaceChart: 'surface',
  surface3DChart: '3-D surface',
};

export interface ChartInfo {
  chartType: ChartType | undefined;
  /** Why the chart can't be filled, phrased to follow "Chart 'x' on slide n ...". */
  problem?: string;
  sample: ChartData;
  /** Worksheet name used in the series formulas (`Sheet1` in most templates). */
  sheetName: string;
  seriesCount: number;
  /** Fewest series the data may have: a combo must keep every chart group non-empty. */
  minSeries: number;
}

function plotArea(doc: Document): Element | undefined {
  return path(doc.documentElement!, NS.c, 'chart', 'plotArea');
}

function typeElements(pa: Element): Element[] {
  return children(pa, NS.c).filter((e) => (e.localName ?? '').endsWith('Chart'));
}

const orderOf = (ser: Element): number => Number(child(ser, NS.c, 'order')?.getAttribute('val') ?? '0');

/** Every series across all chart groups, in the chart's series order. */
function allSeries(groups: Element[]): Element[] {
  return groups.flatMap((g) => children(g, NS.c, 'ser')).sort((a, b) => orderOf(a) - orderOf(b));
}

/** The point cache inside a `c:tx`, `c:cat` or `c:val` element. */
function cacheOf(container: Element | undefined): Element | undefined {
  if (!container) return undefined;
  const str = child(container, NS.c, 'strRef');
  if (str) return child(str, NS.c, 'strCache');
  const num = child(container, NS.c, 'numRef');
  if (num) return child(num, NS.c, 'numCache');
  const multi = child(container, NS.c, 'multiLvlStrRef');
  if (multi) {
    const cache = child(multi, NS.c, 'multiLvlStrCache');
    return cache && child(cache, NS.c, 'lvl');
  }
  return child(container, NS.c, 'strLit') ?? child(container, NS.c, 'numLit');
}

function readPoints(cache: Element | undefined, ptCountFrom?: Element): (string | null)[] {
  if (!cache) return [];
  const countEl = child(ptCountFrom ?? cache, NS.c, 'ptCount');
  const pts = children(cache, NS.c, 'pt');
  const n = countEl ? Number(countEl.getAttribute('val') ?? '0') : pts.length;
  const out: (string | null)[] = Array.from({ length: n }, () => null);
  for (const pt of pts) {
    const idx = Number(pt.getAttribute('idx') ?? '-1');
    const v = child(pt, NS.c, 'v');
    if (idx >= 0 && idx < n && v) out[idx] = textOf(v);
  }
  return out;
}

function sheetFromFormula(f: string | undefined): string | undefined {
  if (!f) return undefined;
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec(f.trim());
  if (!m) return undefined;
  return m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
}

function formulaOf(container: Element | undefined): string | undefined {
  if (!container) return undefined;
  for (const ref of ['numRef', 'strRef', 'multiLvlStrRef']) {
    const f = path(container, NS.c, ref, 'f');
    if (f) return textOf(f);
  }
  return undefined;
}

export function inspectChart(doc: Document): ChartInfo {
  const info: ChartInfo = { chartType: undefined, sample: { categories: [], series: [] }, sheetName: 'Sheet1', seriesCount: 0, minSeries: 1 };
  const pa = plotArea(doc);
  const types = pa ? typeElements(pa) : [];
  const first = types[0];
  if (!first) {
    info.problem = 'is not a chart DeckForge can read';
    return info;
  }

  const sers = allSeries(types);
  info.seriesCount = sers.length;
  const firstSer = sers[0];
  if (firstSer) {
    const multi = path(firstSer, NS.c, 'cat', 'multiLvlStrRef', 'multiLvlStrCache');
    info.sample.categories = readPoints(cacheOf(child(firstSer, NS.c, 'cat')), multi);
    info.sheetName = sheetFromFormula(formulaOf(child(firstSer, NS.c, 'val'))) ?? 'Sheet1';
  }
  info.sample.series = sers.map((ser, i) => {
    const tx = child(ser, NS.c, 'tx');
    const name = tx ? (readPoints(cacheOf(tx))[0] ?? textOf(child(tx, NS.c, 'v') ?? tx)) : `Series ${i + 1}`;
    const values = readPoints(cacheOf(child(ser, NS.c, 'val'))).map((v) => (v === null || v === '' ? null : Number(v)));
    return { name, values };
  });

  const bad = types.find((g) => !SUPPORTED[g.localName ?? '']);
  if (bad) {
    const local = bad.localName ?? '';
    const label = TYPE_LABELS[local] ?? local.replace(/Chart$/, '');
    info.problem =
      (types.length > 1 ? `contains a ${label} chart part` : `is a ${label} chart`) +
      '. DeckForge can fill bar, column, line, area, pie and doughnut charts (and combinations of them); change the chart type in PowerPoint';
    return info;
  }
  if (sers.length === 0) {
    info.problem = 'has no data series. Add at least one series in PowerPoint (Chart Design > Edit Data)';
    return info;
  }
  const kinds = new Set(types.map((g) => SUPPORTED[g.localName ?? '']!));
  info.chartType = types.length > 1 ? 'combo' : [...kinds][0]!;
  // Each group's first series must survive trimming, so the data needs at least up to the last of those.
  const filled = types.filter((g) => children(g, NS.c, 'ser').length > 0);
  info.minSeries = Math.max(1, ...filled.map((g) => Math.min(...children(g, NS.c, 'ser').map((s) => sers.indexOf(s))) + 1));
  return info;
}

export function columnLetter(n: number): string {
  // 1 -> A, 26 -> Z, 27 -> AA
  let s = '';
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function el(ctx: Element, local: string, attrs: Record<string, string> = {}, text?: string): Element {
  const e = create(ctx, NS.c, local);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  if (text !== undefined) setText(e, text);
  return e;
}

function strRef(ctx: Element, f: string, values: string[]): Element {
  const ref = el(ctx, 'strRef');
  ref.appendChild(el(ctx, 'f', {}, f));
  const cache = el(ctx, 'strCache');
  cache.appendChild(el(ctx, 'ptCount', { val: String(values.length) }));
  values.forEach((v, idx) => {
    const pt = el(ctx, 'pt', { idx: String(idx) });
    pt.appendChild(el(ctx, 'v', {}, v));
    cache.appendChild(pt);
  });
  ref.appendChild(cache);
  return ref;
}

function numRef(ctx: Element, f: string, formatCode: string, values: (number | null)[]): Element {
  const ref = el(ctx, 'numRef');
  ref.appendChild(el(ctx, 'f', {}, f));
  const cache = el(ctx, 'numCache');
  cache.appendChild(el(ctx, 'formatCode', {}, formatCode));
  cache.appendChild(el(ctx, 'ptCount', { val: String(values.length) }));
  values.forEach((v, idx) => {
    if (v === null || !Number.isFinite(v)) return;
    const pt = el(ctx, 'pt', { idx: String(idx) });
    pt.appendChild(el(ctx, 'v', {}, String(v)));
    cache.appendChild(pt);
  });
  ref.appendChild(cache);
  return ref;
}

/** Replaces the contents of `c:<local>` in a series, creating it before `before` if missing. */
function replaceSeriesPart(ser: Element, local: string, content: Element, before: string[]): void {
  let part = child(ser, NS.c, local);
  if (!part) {
    part = el(ser, local);
    const ref = before.map((b) => child(ser, NS.c, b)).find((x) => x !== undefined);
    ser.insertBefore(part, ref ?? null);
  } else {
    removeChildren(part);
  }
  part.appendChild(content);
}

/** Per-point overrides (`c:dPt`, `c:dLbl`) past the new point count would point at nothing. */
function pruneByIdx(parent: Element, local: string, count: number): void {
  for (const e of children(parent, NS.c, local)) {
    const idx = Number(child(e, NS.c, 'idx')?.getAttribute('val') ?? '0');
    if (idx >= count) removeNode(e);
  }
}

/**
 * Rewrites a chart's series (names, categories, values, caches and range formulas) for `data`.
 * The caller has already checked the chart is supported and `data.series.length <= seriesCount`.
 * `categories` are pre-rendered strings. Returns each series' number format code.
 */
export function fillChart(doc: Document, info: ChartInfo, categories: string[], data: ChartData): string[] {
  const sers = allSeries(typeElements(plotArea(doc)!));
  for (const extra of sers.slice(data.series.length)) removeNode(extra);

  const n = categories.length;
  const sheet = quoteSheet(info.sheetName);
  const lastRow = Math.max(n + 1, 2);

  const formatCodes: string[] = [];
  data.series.forEach((series, i) => {
    const ser = sers[i]!;
    const col = columnLetter(i + 2);
    const val = child(ser, NS.c, 'val');
    const fcEl = val ? descendants(val, NS.c, 'formatCode')[0] : undefined;
    const formatCode = (fcEl && textOf(fcEl)) || 'General';
    formatCodes.push(formatCode);

    replaceSeriesPart(ser, 'tx', strRef(ser, `${sheet}!$${col}$1`, [series.name]), ['spPr', 'invertIfNegative', 'marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val']);
    replaceSeriesPart(ser, 'cat', strRef(ser, `${sheet}!$A$2:$A$${lastRow}`, categories), ['val']);
    const values = Array.from({ length: n }, (_, k) => series.values[k] ?? null);
    replaceSeriesPart(ser, 'val', numRef(ser, `${sheet}!$${col}$2:$${col}$${lastRow}`, formatCode, values), ['smooth', 'shape', 'extLst']);

    pruneByIdx(ser, 'dPt', n);
    const dLbls = child(ser, NS.c, 'dLbls');
    if (dLbls) pruneByIdx(dLbls, 'dLbl', n);
  });
  return formatCodes;
}

/** Relationship id of the chart's embedded workbook, if it has one. */
export function externalDataRelId(doc: Document): string | undefined {
  const ext = child(doc.documentElement!, NS.c, 'externalData');
  return ext?.getAttributeNS(NS.r, 'id') ?? undefined;
}
