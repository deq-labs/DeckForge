import { OpcPackage } from './opc.ts';
import { NS, REL, child, descendants, path, textOf } from './xml.ts';
import type { Document, Element } from './xml.ts';
import { SLOT_NAME_RE, findTags, paragraphLines, parseTag, removeLines, replaceTags } from './tags.ts';
import { fillTable, normalizeHeader, readTable } from './table.ts';
import { externalDataRelId, fillChart, inspectChart } from './chart.ts';
import type { ChartInfo } from './chart.ts';
import { buildChartWorkbook } from './workbook.ts';
import { formatScalar, normalizeEntry, renderField, validateFormat } from './format.ts';
import type { FormatContext } from './format.ts';
import { deleteSlide, duplicateSlide, listSlides, notesPathOf, presentationPath } from './slides.ts';
import type {
  ChartSlot, DeckData, FillOptions, Issue, Manifest, RepeatSlot, Scalar, SlideRule, TableSlot, TextField, TextSlot, ValueEntry,
} from './types.ts';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const REL_MASTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const REL_DIAGRAM_DATA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData';
const MAX_SHEET_NAME = 31;
const DEFAULT_SLIDE_HEIGHT = 6858000;

const DIRECTIVE_RE = /^\s*#(repeat|show-if|hide-if)(?:\s+(\S+))?\s*$/i;
const LOOKS_LIKE_DIRECTIVE = /^\s*#(repeat|show|hide|skip)/i;

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

type Lookup = Map<string, ValueEntry | Scalar>;

interface SlideInfo {
  /** Template slide number; copies keep their source's number. */
  number: number;
  path: string;
  /** Paragraphs whose tags get filled: slide, notes (minus directives), chart titles. */
  paragraphs: Element[];
  directiveParagraphs: Element[];
  repeat: string | undefined;
  showIf: string[];
  hideIf: string[];
}
interface LayoutInfo {
  name: string;
  paragraphs: Element[];
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
  slides: SlideInfo[];
  layouts: LayoutInfo[];
  tables: Map<string, TableLoc[]>;
  charts: Map<string, ChartLoc[]>;
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
    const analysis = await analyze(await OpcPackage.load(bytes), {});
    return new Template(bytes, analysis.manifest);
  }

  /** What the template can be filled with, plus problems found in the template itself. */
  slots(): Manifest {
    return structuredClone(this.manifest);
  }

  /** Fills a fresh copy of the template. The template itself is unchanged and can be filled again. */
  async fill(data: DeckData, options: FillOptions = {}): Promise<Deck> {
    const pkg = await OpcPackage.load(this.bytes);
    const first = await analyze(pkg, {});
    const ctx: FormatContext = { locale: options.locale ?? 'en-US', currency: options.currency ?? 'USD' };
    const onMissing = options.onMissing ?? 'error';
    const missingSeverity = onMissing === 'error' ? 'error' : 'warning';
    const issues: Issue[] = [...first.manifest.issues];
    const seen = new Set<string>();
    const report = (issue: Issue): void => {
      if (seen.has(issue.message)) return;
      seen.add(issue.message);
      issues.push(issue);
    };

    const values = lowerKeys(data.values);
    const tables = lowerKeys(data.tables);
    const charts = lowerKeys(data.charts);
    const rowSets = new Map([...lowerKeys(data.rows)].map(([k, rows]) => [k, rows.map((r) => lowerKeys(r))]));

    // Values with an unknown format are reported once and their tags left in place.
    const badFormat = new Set<string>();
    const checkFormat = (key: string, raw: ValueEntry | Scalar, where: string): void => {
      const problem = validateFormat(normalizeEntry(raw)?.format);
      if (problem && !badFormat.has(key)) {
        badFormat.add(key);
        report({ severity: 'error', code: 'bad-value', slot: key, message: `${where}: ${problem}` });
      }
    };
    for (const [key, raw] of values) checkFormat(key, raw, `Value "${key}"`);
    for (const [source, rows] of rowSets) for (const row of rows) for (const [key, raw] of row) checkFormat(key, raw, `Rows '${source}', "${key}"`);

    // 1. Slide rules: repeat, then show-if / hide-if on every resulting copy.
    const origin = new Map<string, number>();
    const overlays = new Map<string, Lookup>();
    for (const s of first.slides) origin.set(s.path, s.number);
    for (const s of first.slides) {
      let copies: { path: string; overlay?: Lookup }[] = [{ path: s.path }];
      if (s.repeat) {
        const rows = rowSets.get(s.repeat.toLowerCase());
        if (!rows) {
          report({ severity: missingSeverity, code: 'missing-rows', slide: s.number, slot: s.repeat, message: `No rows for "#repeat ${s.repeat}" (slide ${s.number}).` });
        } else {
          copies = [];
          let after = s.path;
          for (const [i, row] of rows.entries()) {
            const copyPath = i === 0 ? s.path : await duplicateSlide(pkg, s.path, after);
            origin.set(copyPath, s.number);
            copies.push({ path: copyPath, overlay: row });
            after = copyPath;
          }
          if (rows.length === 0) await deleteSlide(pkg, s.path);
        }
      }
      for (const c of copies) {
        if (c.overlay) overlays.set(c.path, c.overlay);
        const look = (key: string): ValueEntry | Scalar | undefined => {
          const k = key.toLowerCase();
          return c.overlay?.has(k) ? c.overlay.get(k) : values.get(k);
        };
        let visible = true;
        for (const key of s.showIf) {
          const v = look(key);
          if (v === undefined) {
            report({ severity: 'warning', code: 'slide-rule', slide: s.number, slot: key, message: `Slide ${s.number}: "#show-if ${key}" has no data for ${key}, so the slide is left out.` });
          }
          if (!truthy(v)) visible = false;
        }
        for (const key of s.hideIf) if (truthy(look(key))) visible = false;
        if (!visible) await deleteSlide(pkg, c.path);
      }
    }

    // 2. Re-read the deck as it now stands.
    const a = await analyze(pkg, { origin, fillPass: true });

    // 3. Text tags (before tables, so data written into table cells is never read as a tag).
    const missing = new Map<string, { name: string; slides: Set<number>; layouts: Set<string> }>();
    const resolver = (overlay: Lookup | undefined, where: { slide?: number; layout?: string }) => (inner: string): string | undefined => {
      const tag = parseTag(inner);
      if (!tag.ok) return undefined;
      const key = tag.name.toLowerCase();
      const raw = overlay?.has(key) ? overlay.get(key) : values.get(key);
      const entry = normalizeEntry(raw);
      if (!entry) {
        const m = missing.get(key) ?? { name: tag.name, slides: new Set<number>(), layouts: new Set<string>() };
        if (where.slide !== undefined) m.slides.add(where.slide);
        if (where.layout !== undefined) m.layouts.add(where.layout);
        missing.set(key, m);
        return onMissing === 'blank' ? '' : undefined;
      }
      if (badFormat.has(key)) return undefined;
      const r = renderField(entry, tag.field, ctx);
      if ('error' in r) {
        const place = where.slide !== undefined ? `Slide ${where.slide}` : `Layout '${where.layout}'`;
        report({ severity: 'error', code: 'bad-value', ...(where.slide !== undefined ? { slide: where.slide } : {}), slot: tag.name, message: `${place}: {{${inner.trim()}}} ${r.error}.` });
        return undefined;
      }
      return r.text;
    };
    for (const slide of a.slides) {
      const resolve = resolver(overlays.get(slide.path), { slide: slide.number });
      for (const p of slide.paragraphs) replaceTags(p, resolve);
      for (const p of slide.directiveParagraphs) removeLines(p, (line) => isDirective(line) !== undefined);
    }
    for (const layout of a.layouts) {
      const resolve = resolver(undefined, { layout: layout.name });
      for (const p of layout.paragraphs) replaceTags(p, resolve);
    }
    const order = [...first.manifest.text.map((t) => t.name), ...first.manifest.repeats.flatMap((r) => r.tags.map((t) => t.name))].map((n) => n.toLowerCase());
    const missingKeys = [...missing.keys()].sort((x, y) => rank(order, x) - rank(order, y));
    for (const key of missingKeys) {
      const m = missing.get(key)!;
      const slides = [...m.slides].sort((x, y) => x - y);
      const places = [
        ...(slides.length ? [`slide ${slides.join(', ')}`] : []),
        ...[...m.layouts].map((l) => `layout '${l}'`),
      ];
      report({
        severity: missingSeverity,
        code: 'missing-value',
        slot: m.name,
        ...(slides.length === 1 && m.layouts.size === 0 ? { slide: slides[0]! } : {}),
        message: `No data for {{${m.name}}} (${places.join('; ')}).`,
      });
    }

    // 4. Tables (every copy of a repeated slide gets the same table data).
    for (const [key, locs] of a.tables) {
      const td = tables.get(key);
      for (const loc of locs) {
        const { name, slide } = loc.slot;
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
          rows.push(
            r.map((cell) => {
              const e = normalizeEntry(cell)!;
              const problem = validateFormat(e.format);
              if (problem) {
                badCell ??= problem;
                return String(e.value ?? '');
              }
              return formatScalar(e.value, e.format, ctx);
            }),
          );
        }
        if (badCell) {
          report({ severity: 'error', code: 'bad-value', slide, slot: name, message: `Table '${name}' on slide ${slide}: ${badCell}` });
          continue;
        }
        fillTable(loc.frame, loc.tbl, columnMap, rows);
        const bottom = frameBottom(loc.frame);
        if (bottom !== undefined && bottom > a.manifest.slideHeight) {
          report({
            severity: 'warning',
            code: 'overflow',
            slide,
            slot: name,
            message: `Table '${name}' on slide ${slide} now runs past the bottom of the slide (${rows.length} rows). Make the rows or font smaller in the template, or send fewer rows.`,
          });
        }
      }
    }

    // 5. Charts.
    for (const [key, locs] of a.charts) {
      const cd = charts.get(key);
      for (const loc of locs) {
        const { name, slide } = loc.slot;
        if (!cd && loc.info.problem) continue;
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
        const need = loc.info.minSeries;
        if (want < need || want > have) {
          const s = (n: number): string => `${n} series column${n === 1 ? '' : 's'}`;
          const message =
            want === 0
              ? `Chart '${name}' on slide ${slide} needs at least one series column in the data.`
              : want > have
                ? `Chart '${name}' on slide ${slide} has ${have} series in the template but the data has ${want}. ` +
                  `Remove ${s(want - have)} from the data, or add series to the chart in PowerPoint (Chart Design > Edit Data).`
                : `Chart '${name}' on slide ${slide} combines chart types and needs at least ${need} series so each part keeps one; the data has ${want}.`;
          report({ severity: 'error', code: 'chart-series', slide, slot: name, message });
          continue;
        }
        if ((loc.info.chartType === 'pie' || loc.info.chartType === 'doughnut') && cd.series.some((sr) => sr.values.some((v) => v !== null && v < 0))) {
          report({ severity: 'warning', code: 'bad-value', slide, slot: name, message: `Chart '${name}' on slide ${slide} is a ${loc.info.chartType} chart but has negative values; PowerPoint draws them as positive slices.` });
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
    }

    // 6. Data that matches nothing in the template.
    for (const [key] of tables) {
      if (!first.tables.has(key)) report({ severity: 'warning', code: 'unused-data', slot: key, message: `The data has a table '${key}', but the template has no table with that name.` });
    }
    for (const [key] of charts) {
      if (!first.charts.has(key)) report({ severity: 'warning', code: 'unused-data', slot: key, message: `The data has a chart '${key}', but the template has no chart with that name.` });
    }
    const repeatNames = new Set(first.manifest.repeats.map((r) => r.name.toLowerCase()));
    for (const [key] of rowSets) {
      if (!repeatNames.has(key)) report({ severity: 'warning', code: 'unused-data', slot: key, message: `The data has rows '${key}', but no slide in the template says "#repeat ${key}".` });
    }

    if (issues.some((i) => i.severity === 'error')) throw new FillError(issues);
    await pkg.collectGarbage();
    return new Deck(await pkg.save(), issues);
  }
}

function isDirective(line: string): { kind: SlideRule['kind']; target: string } | undefined {
  const m = DIRECTIVE_RE.exec(line);
  if (!m || !m[2] || !SLOT_NAME_RE.test(m[2])) return undefined;
  return { kind: m[1]!.toLowerCase() as SlideRule['kind'], target: m[2] };
}

function rank(order: string[], key: string): number {
  const i = order.indexOf(key);
  return i < 0 ? order.length : i;
}

function lowerKeys<T>(rec: Record<string, T> | undefined): Map<string, T> {
  const m = new Map<string, T>();
  for (const [k, v] of Object.entries(rec ?? {})) m.set(k.toLowerCase(), v);
  return m;
}

/** Empty, 0, "no", "false", "none" and "n" are false; anything else present is true. */
export function truthy(v: ValueEntry | Scalar | undefined): boolean {
  const e = normalizeEntry(v);
  if (!e) return false;
  const x = e.value;
  if (x === null) return false;
  if (typeof x === 'number') return x !== 0;
  if (typeof x === 'boolean') return x;
  if (x instanceof Date) return true;
  return !['', '0', 'no', 'n', 'false', 'none'].includes(x.trim().toLowerCase());
}

function categoryText(v: Scalar, ctx: FormatContext): string {
  if (v === null) return '';
  if (v instanceof Date) return formatScalar(v, 'date', ctx);
  return String(v);
}

function frameBottom(frame: Element): number | undefined {
  const xfrm = path(frame, NS.p, 'xfrm');
  const off = xfrm && child(xfrm, NS.a, 'off');
  const ext = xfrm && child(xfrm, NS.a, 'ext');
  if (!off || !ext) return undefined;
  return Number(off.getAttribute('y') ?? '0') + Number(ext.getAttribute('cy') ?? '0');
}

function insideFallback(el: Element): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    if ((n as Element).localName === 'Fallback') return true;
  }
  return false;
}

function isPlaceholder(sp: Element): boolean {
  return path(sp, NS.p, 'nvSpPr', 'nvPr', 'ph') !== undefined;
}

function notesBodyParagraphs(notes: Document): Element[] {
  const out: Element[] = [];
  for (const sp of descendants(notes, NS.p, 'sp')) {
    const ph = path(sp, NS.p, 'nvSpPr', 'nvPr', 'ph');
    if (ph?.getAttribute('type') === 'body') out.push(...descendants(sp, NS.a, 'p'));
  }
  return out;
}

interface AnalyzeOptions {
  /** Output slide path -> template slide number, for decks with copied slides. */
  origin?: Map<string, number>;
  /** Second pass over a deck with copies: repeated names are expected, rule problems already reported. */
  fillPass?: boolean;
}

async function analyze(pkg: OpcPackage, opts: AnalyzeOptions): Promise<Analysis> {
  const issues: Issue[] = [];
  const text = new Map<string, TextSlot>();
  const repeats = new Map<string, RepeatSlot>();
  const rules: SlideRule[] = [];
  const tables = new Map<string, TableLoc[]>();
  const charts = new Map<string, ChartLoc[]>();
  const slides: SlideInfo[] = [];
  const invalidSeen = new Set<string>();
  const layoutPaths: string[] = [];
  const warn = (issue: Omit<Issue, 'severity'>): void => {
    if (!opts.fillPass) issues.push({ severity: 'warning', ...issue });
  };

  const pres = (await pkg.xml(await presentationPath(pkg))).documentElement!;
  const sldSz = descendants(pres, NS.p, 'sldSz')[0];
  const slideHeight = Number(sldSz?.getAttribute('cy') ?? DEFAULT_SLIDE_HEIGHT);

  const validTag = (n: number, inner: string): boolean => {
    const tag = parseTag(inner);
    if (tag.ok) return true;
    const k = `${n}|${inner}`;
    if (!invalidSeen.has(k)) {
      invalidSeen.add(k);
      warn({ code: 'invalid-tag', slide: n, message: `Slide ${n}: "{{${inner}}}" looks like a tag but will not be filled: ${tag.reason}.` });
    }
    return false;
  };

  for (const [i, ref] of (await listSlides(pkg)).entries()) {
    const n = opts.origin?.get(ref.path) ?? i + 1;
    const doc = await pkg.xml(ref.path);
    const slideRels = await pkg.rels(ref.path);
    const info: SlideInfo = { number: n, path: ref.path, paragraphs: descendants(doc, NS.a, 'p'), directiveParagraphs: [], repeat: undefined, showIf: [], hideIf: [] };

    // Speaker notes: directives, and tags to fill.
    const notesPath = await notesPathOf(pkg, ref.path);
    if (notesPath) {
      for (const p of notesBodyParagraphs(await pkg.xml(notesPath))) {
        let hasDirective = false;
        for (const line of paragraphLines(p)) {
          const m = isDirective(line);
          if (m) {
            hasDirective = true;
            if (!opts.fillPass) rules.push({ slide: n, kind: m.kind, target: m.target });
            if (m.kind === 'repeat') {
              if (info.repeat && !opts.fillPass) {
                issues.push({ severity: 'error', code: 'slide-rule', slide: n, message: `Slide ${n} has more than one #repeat line in its notes. A slide can repeat over one sheet only.` });
              }
              info.repeat ??= m.target;
            } else if (m.kind === 'show-if') info.showIf.push(m.target);
            else info.hideIf.push(m.target);
          } else if (LOOKS_LIKE_DIRECTIVE.test(line)) {
            warn({ code: 'slide-rule', slide: n, message: `Slide ${n}: the notes line "${line.trim()}" looks like a DeckForge instruction but isn't one. Use "#repeat <sheet>", "#show-if <key>" or "#hide-if <key>".` });
          }
        }
        if (hasDirective) info.directiveParagraphs.push(p);
        info.paragraphs.push(p);
      }
    }

    // Charts (titles get tags too), layouts, SmartArt.
    for (const r of slideRels) {
      if (r.external || !pkg.has(r.target)) continue;
      if (r.type === REL.chart) info.paragraphs.push(...descendants(await pkg.xml(r.target), NS.a, 'p'));
      if (r.type === REL_LAYOUT && !layoutPaths.includes(r.target)) layoutPaths.push(r.target);
      if (r.type === REL_DIAGRAM_DATA && /\{\{[^{}]*\}\}/.test(textOf((await pkg.xml(r.target)).documentElement!))) {
        warn({ code: 'misplaced-tag', slide: n, message: `Slide ${n}: a SmartArt graphic contains a {{tag}}. SmartArt is not supported; convert it to shapes (right-click > Convert to Shapes).` });
      }
    }
    slides.push(info);

    for (const p of info.paragraphs) {
      for (const inner of findTags(p)) {
        if (!validTag(n, inner)) continue;
        const tag = parseTag(inner);
        if (!tag.ok) continue;
        const key = tag.name.toLowerCase();
        if (info.repeat) {
          const rk = info.repeat.toLowerCase();
          const rs = repeats.get(rk) ?? { kind: 'repeat' as const, name: info.repeat, slides: [] as number[], tags: [] as { name: string; fields: TextField[] }[] };
          let t = rs.tags.find((x) => x.name.toLowerCase() === key);
          if (!t) rs.tags.push((t = { name: tag.name, fields: [] }));
          if (!t.fields.includes(tag.field)) t.fields.push(tag.field);
          repeats.set(rk, rs);
        } else {
          const slot = text.get(key) ?? { kind: 'text' as const, name: tag.name, fields: [] as TextField[], slides: [] as number[], layouts: [] as string[] };
          if (!slot.fields.includes(tag.field)) slot.fields.push(tag.field);
          if (!slot.slides.includes(n)) slot.slides.push(n);
          text.set(key, slot);
        }
      }
    }
    if (info.repeat) {
      const rk = info.repeat.toLowerCase();
      const rs = repeats.get(rk) ?? { kind: 'repeat' as const, name: info.repeat, slides: [] as number[], tags: [] as { name: string; fields: TextField[] }[] };
      if (!rs.slides.includes(n)) rs.slides.push(n);
      repeats.set(rk, rs);
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
        warn({ code: 'invalid-tag', slide: n, slot: name, message: `${what} '${name}' on slide ${n} has a name longer than ${MAX_SHEET_NAME} characters, so no Excel sheet can match it. Shorten it in the Selection Pane.` });
        continue;
      }
      const key = name.toLowerCase();
      const prev = tables.get(key)?.[0]?.slot ?? charts.get(key)?.[0]?.slot;
      if (prev && !opts.fillPass) {
        issues.push({ severity: 'error', code: 'duplicate-slot', slide: n, slot: name, message: `Two tables or charts are named '${name}' (slides ${prev.slide} and ${n}). Give each one its own name in the Selection Pane.` });
        continue;
      }

      if (tbl) {
        const t = readTable(tbl);
        const loc: TableLoc = { slot: { kind: 'table', name, slide: n, columns: t.columns, sampleRows: t.sampleRows }, frame, tbl, problem: t.problem };
        tables.set(key, [...(tables.get(key) ?? []), loc]);
        if (t.problem) warn({ code: 'unsupported', slide: n, slot: name, message: `Table '${name}' on slide ${n} ${t.problem}.` });
      } else {
        const rid = chartEl!.getAttributeNS(NS.r, 'id') ?? '';
        const rel = await pkg.relById(ref.path, rid);
        if (!rel || rel.external || !pkg.has(rel.target)) {
          warn({ code: 'unsupported', slide: n, slot: name, message: `Chart '${name}' on slide ${n} links to data DeckForge cannot find inside the file.` });
          continue;
        }
        const chartDoc = await pkg.xml(rel.target);
        const ci = inspectChart(chartDoc);
        const loc: ChartLoc = { slot: { kind: 'chart', name, slide: n, chartType: ci.chartType, sample: ci.sample }, partPath: rel.target, doc: chartDoc, info: ci };
        charts.set(key, [...(charts.get(key) ?? []), loc]);
        if (ci.problem) warn({ code: 'unsupported', slide: n, slot: name, message: `Chart '${name}' on slide ${n} ${ci.problem}.` });
      }
    }
  }

  // Slide layouts and their masters: non-placeholder shapes get filled; placeholders never show.
  const layouts: LayoutInfo[] = [];
  const masterPaths: string[] = [];
  for (const lp of layoutPaths) {
    for (const r of await pkg.rels(lp)) if (r.type === REL_MASTER && !r.external && !masterPaths.includes(r.target)) masterPaths.push(r.target);
  }
  for (const lp of [...layoutPaths, ...masterPaths]) {
    if (!pkg.has(lp)) continue;
    const doc = await pkg.xml(lp);
    const isMaster = masterPaths.includes(lp);
    const name = isMaster ? 'Slide Master' : (descendants(doc, NS.p, 'cSld')[0]?.getAttribute('name') || lp.split('/').pop()!);
    const info: LayoutInfo = { name, paragraphs: [] };
    for (const sp of descendants(doc, NS.p, 'sp')) {
      const paras = descendants(sp, NS.a, 'p');
      const tags = paras.flatMap((p) => findTags(p)).filter((t) => parseTag(t).ok);
      if (tags.length === 0) continue;
      if (isPlaceholder(sp)) {
        warn({
          code: 'misplaced-tag',
          message: `${isMaster ? 'The slide master' : `Slide layout '${name}'`} has {{${tags[0]!.trim()}}} inside a placeholder. Placeholder text on a layout is only a prompt and never appears on slides, so it is not filled. Put the tag on the slides, or in a normal text box on the layout.`,
        });
        continue;
      }
      info.paragraphs.push(...paras);
      for (const inner of tags) {
        const tag = parseTag(inner);
        if (!tag.ok) continue;
        const key = tag.name.toLowerCase();
        const slot = text.get(key) ?? { kind: 'text' as const, name: tag.name, fields: [] as TextField[], slides: [] as number[], layouts: [] as string[] };
        if (!slot.fields.includes(tag.field)) slot.fields.push(tag.field);
        if (!slot.layouts.includes(name)) slot.layouts.push(name);
        text.set(key, slot);
      }
    }
    if (info.paragraphs.length) layouts.push(info);
  }

  if (!opts.fillPass) {
    for (const r of repeats.values()) {
      const k = r.name.toLowerCase();
      if (tables.has(k) || charts.has(k)) {
        issues.push({ severity: 'error', code: 'duplicate-slot', slot: r.name, message: `"#repeat ${r.name}" (slide ${r.slides.join(', ')}) uses the same name as a table or chart. Use a different sheet name for the rows.` });
      }
    }
  }

  return {
    manifest: {
      slideCount: slides.length,
      slideHeight,
      text: [...text.values()],
      tables: [...tables.values()].map((t) => t[0]!.slot),
      charts: [...charts.values()].map((c) => c[0]!.slot),
      repeats: [...repeats.values()],
      rules,
      issues,
    },
    slides,
    layouts,
    tables,
    charts,
  };
}
