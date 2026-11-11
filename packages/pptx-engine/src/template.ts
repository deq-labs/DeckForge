import { OpcPackage } from './opc.ts';
import { NS, REL, child, descendants, path } from './xml.ts';
import type { Document, Element } from './xml.ts';
import { SLOT_NAME_RE, findTags, parseTag, replaceTags } from './tags.ts';
import { fillTable, normalizeHeader, readTable } from './table.ts';
import { externalDataRelId, fillChart, inspectChart } from './chart.ts';
import type { ChartInfo } from './chart.ts';
import { buildChartWorkbook } from './workbook.ts';
import { formatScalar, normalizeEntry, renderField, validateFormat } from './format.ts';
import type { FormatContext } from './format.ts';
import type {
  ChartSlot, DeckData, FillOptions, Issue, Manifest, Scalar, TableSlot, TextField, TextSlot,
} from './types.ts';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const MAX_SHEET_NAME = 31;

/** Thrown by `Template.fill` when the deck can't be produced. `issues` explains why. */
export class FillError extends Error {
  readonly issues: Issue[];
  constructor(issues: Issue[]) {
    const errors = issues.filter((i) => i.severity === 'error');
    super(`The deck could not be filled (${errors.length} problem${errors.length === 1 ? '' : 's'}):\n- ${errors.map((i) => i.message).join('\n- ')}`);
    this.name = 'FillError';
    this.issues = issues;
  }
}

export class Deck {
  readonly issues: Issue[];
  private readonly bytes: Uint8Array;
  constructor(bytes: Uint8Array, issues: Issue[]) {
    this.bytes = bytes;
    this.issues = issues;
  }
  toBuffer(): Uint8Array {
    return this.bytes;
  }
}

interface SlideRef {
  number: number;
  doc: Document;
}
interface TableLoc {
  slot: TableSlot;
  frame: Element;
  tbl: Element;
  problem: string | undefined;
}
interface ChartLoc {
  slot: ChartSlot;
  partPath: string;
  doc: Document;
  info: ChartInfo;
}
interface Analysis {
  manifest: Manifest;
  slides: SlideRef[];
  tables: Map<string, TableLoc>;
  charts: Map<string, ChartLoc>;
}

export class Template {
  private readonly bytes: Uint8Array;
  private readonly manifest: Manifest;

  private constructor(bytes: Uint8Array, manifest: Manifest) {
    this.bytes = bytes;
    this.manifest = manifest;
  }

  static async load(data: Uint8Array): Promise<Template> {
    const bytes = new Uint8Array(data); // own copy: callers may reuse their buffer
    const analysis = await analyze(await OpcPackage.load(bytes));
    return new Template(bytes, analysis.manifest);
  }

  /** What the template can be filled with, plus problems found in the template itself. */
  slots(): Manifest {
    return structuredClone(this.manifest);
  }

  /** Fills a fresh copy of the template. The template itself is unchanged and can be filled again. */
  async fill(data: DeckData, options: FillOptions = {}): Promise<Deck> {
    const pkg = await OpcPackage.load(this.bytes);
    const a = await analyze(pkg);
    const ctx: FormatContext = { locale: options.locale ?? 'en-US', currency: options.currency ?? 'USD' };
    const onMissing = options.onMissing ?? 'error';
    const missingSeverity = onMissing === 'error' ? 'error' : 'warning';
    const issues: Issue[] = [...a.manifest.issues];
    const seen = new Set<string>();
    const report = (issue: Issue): void => {
      if (seen.has(issue.message)) return;
      seen.add(issue.message);
      issues.push(issue);
    };

    const values = lowerKeys(data.values);
    const tables = lowerKeys(data.tables);
    const charts = lowerKeys(data.charts);

    // Values with an unknown format are reported once and their tags left in place.
    const badFormat = new Set<string>();
    for (const [key, raw] of values) {
      const entry = normalizeEntry(raw);
      const problem = validateFormat(entry?.format);
      if (problem) {
        badFormat.add(key);
        report({ severity: 'error', code: 'bad-value', slot: key, message: `Value "${key}": ${problem}` });
      }
    }

    // 1. Text tags (before tables, so data written into table cells is never read as a tag).
    const missing = new Map<string, { name: string; slides: Set<number> }>();
    for (const slide of a.slides) {
      for (const p of descendants(slide.doc, NS.a, 'p')) {
        replaceTags(p, (inner) => {
          const tag = parseTag(inner);
          if (!tag.ok) return undefined;
          const key = tag.name.toLowerCase();
          const entry = normalizeEntry(values.get(key));
          if (!entry) {
            const m = missing.get(key) ?? { name: tag.name, slides: new Set<number>() };
            m.slides.add(slide.number);
            missing.set(key, m);
            return onMissing === 'blank' ? '' : undefined;
          }
          if (badFormat.has(key)) return undefined;
          const r = renderField(entry, tag.field, ctx);
          if ('error' in r) {
            report({ severity: 'error', code: 'bad-value', slide: slide.number, slot: tag.name, message: `Slide ${slide.number}: {{${inner.trim()}}} ${r.error}.` });
            return undefined;
          }
          return r.text;
        });
      }
    }
    // Report in reading order (replacement runs right to left within a paragraph).
    for (const slot of a.manifest.text) {
      const m = missing.get(slot.name.toLowerCase());
      if (!m) continue;
      const { name } = m;
      const list = [...m.slides].sort((x, y) => x - y);
      report({
        severity: missingSeverity,
        code: 'missing-value',
        slot: name,
        ...(list.length === 1 ? { slide: list[0]! } : {}),
        message: `No data for {{${name}}} (slide ${list.join(', ')}).`,
      });
    }

    // 2. Tables.
    for (const [key, loc] of a.tables) {
      const { name, slide } = loc.slot;
      const td = tables.get(key);
      if (!td && loc.problem) continue; // already a template warning; nothing asked of it
      if (!td) {
        report({ severity: missingSeverity, code: 'missing-table', slide, slot: name, message: `No data for table '${name}' (slide ${slide}).` });
        continue;
      }
      if (loc.problem) {
        report({ severity: 'error', code: 'unsupported', slide, slot: name, message: `Table '${name}' on slide ${slide} ${loc.problem}.` });
        continue;
      }
      const dataIdx = new Map<string, number>();
      td.columns.forEach((c, i) => {
        const k = normalizeHeader(c);
        if (!dataIdx.has(k)) dataIdx.set(k, i);
      });
      const columnMap = loc.slot.columns.map((h) => (normalizeHeader(h) === '' ? -1 : (dataIdx.get(normalizeHeader(h)) ?? -2)));
      const absent = loc.slot.columns.filter((_, i) => columnMap[i] === -2);
      if (absent.length > 0) {
        report({
          severity: 'error',
          code: 'table-columns',
          slide,
          slot: name,
          message:
            `Table '${name}' on slide ${slide} has the column${absent.length > 1 ? 's' : ''} ${absent.map((c) => `"${c}"`).join(', ')}, ` +
            `but the data has no column with that header. Data columns: ${td.columns.map((c) => `"${c}"`).join(', ') || '(none)'}.`,
        });
        continue;
      }
      const rows: string[][] = [];
      let badCell: string | undefined;
      for (const r of td.rows) {
        rows.push(r.map((cell) => {
          const e = normalizeEntry(cell)!;
          const problem = validateFormat(e.format);
          if (problem) {
            badCell ??= problem;
            return String(e.value ?? '');
          }
          return formatScalar(e.value, e.format, ctx);
        }));
      }
      if (badCell) {
        report({ severity: 'error', code: 'bad-value', slide, slot: name, message: `Table '${name}' on slide ${slide}: ${badCell}` });
        continue;
      }
      fillTable(loc.frame, loc.tbl, columnMap, rows);
    }

    // 3. Charts.
    for (const [key, loc] of a.charts) {
      const { name, slide } = loc.slot;
      const cd = charts.get(key);
      if (!cd && loc.info.problem) continue; // already a template warning; nothing asked of it
      if (!cd) {
        report({ severity: missingSeverity, code: 'missing-chart', slide, slot: name, message: `No data for chart '${name}' (slide ${slide}).` });
        continue;
      }
      if (loc.info.problem) {
        report({ severity: 'error', code: 'unsupported', slide, slot: name, message: `Chart '${name}' on slide ${slide} ${loc.info.problem}.` });
        continue;
      }
      const have = loc.info.seriesCount;
      const want = cd.series.length;
      if (want === 0 || want > have) {
        report({
          severity: 'error',
          code: 'chart-series',
          slide,
          slot: name,
          message:
            want === 0
              ? `Chart '${name}' on slide ${slide} needs at least one series column in the data.`
              : `Chart '${name}' on slide ${slide} has ${have} series in the template but the data has ${want}. ` +
                `Remove ${want - have} series column${want - have > 1 ? 's' : ''} from the data, or add series to the chart in PowerPoint (Chart Design > Edit Data).`,
        });
        continue;
      }
      const categories = cd.categories.map((v) => categoryText(v, ctx));
      const formatCodes = fillChart(loc.doc, loc.info, categories, cd);
      const relId = externalDataRelId(loc.doc);
      const rel = relId ? await pkg.relById(loc.partPath, relId) : undefined;
      if (rel && !rel.external && rel.target.toLowerCase().endsWith('.xlsx')) {
        pkg.setBinary(rel.target, await buildChartWorkbook(loc.info.sheetName, categories, cd, formatCodes));
      } else {
        report({
          severity: 'warning',
          code: 'unsupported',
          slide,
          slot: name,
          message: `Chart '${name}' on slide ${slide} has no embedded workbook, so "Edit Data" in PowerPoint will not show the new numbers. The chart itself is filled.`,
        });
      }
    }

    // 4. Data that matches nothing in the template.
    for (const [key] of tables) {
      if (!a.tables.has(key)) report({ severity: 'warning', code: 'unused-data', slot: key, message: `The data has a table '${key}', but the template has no table with that name.` });
    }
    for (const [key] of charts) {
      if (!a.charts.has(key)) report({ severity: 'warning', code: 'unused-data', slot: key, message: `The data has a chart '${key}', but the template has no chart with that name.` });
    }

    if (issues.some((i) => i.severity === 'error')) throw new FillError(issues);
    return new Deck(await pkg.save(), issues);
  }
}

function lowerKeys<T>(rec: Record<string, T> | undefined): Map<string, T> {
  const m = new Map<string, T>();
  for (const [k, v] of Object.entries(rec ?? {})) m.set(k.toLowerCase(), v);
  return m;
}

function categoryText(v: Scalar, ctx: FormatContext): string {
  if (v === null) return '';
  if (v instanceof Date) return formatScalar(v, 'date', ctx);
  return String(v);
}

async function listSlides(pkg: OpcPackage): Promise<{ number: number; path: string }[]> {
  const rootRels = await pkg.rels('');
  const main = rootRels.find((r) => r.type === REL.officeDocument)?.target ?? 'ppt/presentation.xml';
  const pres = await pkg.xml(main);
  const rels = await pkg.rels(main);
  const out: { number: number; path: string }[] = [];
  for (const sldId of descendants(pres, NS.p, 'sldId')) {
    const rid = sldId.getAttributeNS(NS.r, 'id');
    const rel = rels.find((r) => r.id === rid && r.type === REL.slide);
    if (rel) out.push({ number: out.length + 1, path: rel.target });
  }
  return out;
}

function insideFallback(el: Element): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    if ((n as Element).localName === 'Fallback') return true;
  }
  return false;
}

async function analyze(pkg: OpcPackage): Promise<Analysis> {
  const issues: Issue[] = [];
  const text = new Map<string, TextSlot>();
  const tables = new Map<string, TableLoc>();
  const charts = new Map<string, ChartLoc>();
  const slides: SlideRef[] = [];
  const invalidSeen = new Set<string>();

  for (const ref of await listSlides(pkg)) {
    const doc = await pkg.xml(ref.path);
    const n = ref.number;
    slides.push({ number: n, doc });

    for (const p of descendants(doc, NS.a, 'p')) {
      for (const inner of findTags(p)) {
        const tag = parseTag(inner);
        if (!tag.ok) {
          const k = `${n}|${inner}`;
          if (!invalidSeen.has(k)) {
            invalidSeen.add(k);
            issues.push({ severity: 'warning', code: 'invalid-tag', slide: n, message: `Slide ${n}: "{{${inner}}}" looks like a tag but will not be filled: ${tag.reason}.` });
          }
          continue;
        }
        const key = tag.name.toLowerCase();
        const slot = text.get(key) ?? { kind: 'text' as const, name: tag.name, fields: [] as TextField[], slides: [] as number[] };
        if (!slot.fields.includes(tag.field)) slot.fields.push(tag.field);
        if (!slot.slides.includes(n)) slot.slides.push(n);
        text.set(key, slot);
      }
    }

    for (const frame of descendants(doc, NS.p, 'graphicFrame')) {
      if (insideFallback(frame)) continue;
      const name = path(frame, NS.p, 'nvGraphicFramePr', 'cNvPr')?.getAttribute('name') ?? '';
      if (!SLOT_NAME_RE.test(name)) continue;
      const gd = path(frame, NS.a, 'graphic', 'graphicData');
      if (!gd) continue;
      const tbl = child(gd, NS.a, 'tbl');
      const chartEl = gd.getAttribute('uri') === CHART_URI ? child(gd, NS.c, 'chart') : undefined;
      if (!tbl && !chartEl) continue;

      const what = tbl ? 'Table' : 'Chart';
      if (name.length > MAX_SHEET_NAME) {
        issues.push({ severity: 'warning', code: 'invalid-tag', slide: n, slot: name, message: `${what} '${name}' on slide ${n} has a name longer than ${MAX_SHEET_NAME} characters, so no Excel sheet can match it. Shorten it in the Selection Pane.` });
        continue;
      }
      const key = name.toLowerCase();
      const prev = tables.get(key)?.slot ?? charts.get(key)?.slot;
      if (prev) {
        issues.push({ severity: 'error', code: 'duplicate-slot', slide: n, slot: name, message: `Two tables or charts are named '${name}' (slides ${prev.slide} and ${n}). Give each one its own name in the Selection Pane.` });
        continue;
      }

      if (tbl) {
        const info = readTable(tbl);
        tables.set(key, { slot: { kind: 'table', name, slide: n, columns: info.columns, sampleRows: info.sampleRows }, frame, tbl, problem: info.problem });
        if (info.problem) issues.push({ severity: 'warning', code: 'unsupported', slide: n, slot: name, message: `Table '${name}' on slide ${n} ${info.problem}.` });
      } else {
        const rid = chartEl!.getAttributeNS(NS.r, 'id') ?? '';
        const rel = await pkg.relById(ref.path, rid);
        if (!rel || rel.external || !pkg.has(rel.target)) {
          issues.push({ severity: 'warning', code: 'unsupported', slide: n, slot: name, message: `Chart '${name}' on slide ${n} links to data DeckForge cannot find inside the file.` });
          continue;
        }
        const chartDoc = await pkg.xml(rel.target);
        const info = inspectChart(chartDoc);
        charts.set(key, { slot: { kind: 'chart', name, slide: n, chartType: info.chartType, sample: info.sample }, partPath: rel.target, doc: chartDoc, info });
        if (info.problem) issues.push({ severity: 'warning', code: 'unsupported', slide: n, slot: name, message: `Chart '${name}' on slide ${n} ${info.problem}.` });
      }
    }
  }

  return {
    manifest: {
      slideCount: slides.length,
      text: [...text.values()],
      tables: [...tables.values()].map((t) => t.slot),
      charts: [...charts.values()].map((c) => c.slot),
      issues,
    },
    slides,
    tables,
    charts,
  };
}

