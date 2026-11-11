import { basename, dirname, extname, join } from 'node:path';
import { FillError, Template } from '@deckforge/pptx-engine';
import type { Issue, Manifest, OnMissing } from '@deckforge/pptx-engine';
import { createWorkbook, readWorkbook } from '@deckforge/data-xlsx';
import { UserError, readInput, samePath, writeOutput } from './files.ts';
import { plural, printIssues } from './report.ts';
import type { Io } from './report.ts';

async function loadTemplate(path: string): Promise<Template> {
  const bytes = readInput(path, 'template');
  try {
    return await Template.load(bytes);
  } catch (e) {
    throw new UserError(`${path}: ${(e as Error).message}`);
  }
}

function withSuffix(path: string, suffix: string, ext: string): string {
  return join(dirname(path), `${basename(path, extname(path))}${suffix}${ext}`);
}

// ---------------------------------------------------------------- slots

export interface SlotsOptions {
  json?: boolean;
}

export async function slots(templatePath: string, opts: SlotsOptions, io: Io): Promise<number> {
  const m = (await loadTemplate(templatePath)).slots();
  if (opts.json) {
    io.out(JSON.stringify(m, null, 2));
    return 0;
  }
  io.out(describe(m, templatePath));
  printIssues(m.issues, io);
  return m.issues.some((i) => i.severity === 'error') ? 1 : 0;
}

function describe(m: Manifest, path: string): string {
  const lines: string[] = [`${basename(path)}: ${plural(m.slideCount, 'slide')}`];
  lines.push('', `Text tags (${m.text.length})  -> kpis sheet`);
  for (const t of m.text) {
    const extra = t.fields.filter((f) => f !== 'value');
    lines.push(`  {{${t.name}}}${extra.length ? `  also ${extra.map((f) => `.${f}`).join(' ')}` : ''}  (slide ${t.slides.join(', ')})`);
  }
  lines.push('', `Tables (${m.tables.length})  -> one sheet each`);
  for (const t of m.tables) lines.push(`  ${t.name}  (slide ${t.slide}): ${t.columns.join(' | ')}`);
  lines.push('', `Charts (${m.charts.length})  -> one sheet each`);
  for (const c of m.charts) {
    const kind = c.chartType ?? 'not fillable';
    lines.push(`  ${c.name}  (slide ${c.slide}, ${kind}): ${c.sample.series.length} series — ${c.sample.series.map((s) => s.name).join(', ')}`);
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------- init-data

export interface InitDataOptions {
  output?: string;
  sample?: boolean;
  force?: boolean;
}

export async function initData(templatePath: string, opts: InitDataOptions, io: Io): Promise<number> {
  const m = (await loadTemplate(templatePath)).slots();
  const out = opts.output ?? withSuffix(templatePath, '.data', '.xlsx');
  const bytes = await createWorkbook(m, { sample: opts.sample ?? true });
  writeOutput(out, bytes, { overwrite: opts.force ?? false, what: 'workbook' });
  printIssues(m.issues, io);
  io.out(`Created ${out}`);
  io.out(`  kpis: ${plural(m.text.length, 'row')} to fill in; ${plural(m.tables.length, 'table sheet')}; ${plural(m.charts.filter((c) => c.chartType).length, 'chart sheet')}.`);
  io.out(`Next: fill in the values, then run  deckforge fill "${templatePath}" "${out}"`);
  return 0;
}

// ---------------------------------------------------------------- fill

export interface FillCliOptions {
  output?: string;
  locale?: string;
  currency?: string;
  missing?: string;
}

const MISSING: OnMissing[] = ['error', 'keep', 'blank'];

export async function fill(templatePath: string, dataPath: string, opts: FillCliOptions, io: Io): Promise<number> {
  const missing = (opts.missing ?? 'error') as OnMissing;
  if (!MISSING.includes(missing)) throw new UserError(`--missing must be one of: ${MISSING.join(', ')}.`);
  const out = opts.output ?? withSuffix(templatePath, '.filled', '.pptx');
  if (samePath(out, templatePath)) throw new UserError('The output would overwrite the template. Choose another name with -o.');
  if (samePath(out, dataPath)) throw new UserError('The output would overwrite the data workbook. Choose another name with -o.');

  const tpl = await loadTemplate(templatePath);
  const dataBytes = readInput(dataPath, 'data workbook');
  let read;
  try {
    read = await readWorkbook(dataBytes, tpl.slots());
  } catch (e) {
    throw new UserError(`${dataPath}: ${(e as Error).message}`);
  }
  if (read.issues.some((i) => i.severity === 'error')) {
    printIssues(read.issues, io);
    io.err(`\nThe workbook has problems; no deck was written. Fix them in ${basename(dataPath)} and run again.`);
    return 1;
  }

  let issues: Issue[];
  try {
    const deck = await tpl.fill(read.data, {
      ...(opts.locale ? { locale: opts.locale } : {}),
      ...(opts.currency ? { currency: opts.currency.toUpperCase() } : {}),
      onMissing: missing,
    });
    writeOutput(out, deck.toBuffer(), { overwrite: true, what: 'deck' });
    issues = [...read.issues, ...deck.issues];
  } catch (e) {
    if (!(e instanceof FillError)) throw e;
    printIssues([...read.issues, ...e.issues], io, { blankKeys: read.blankKeys });
    io.err(`\nNo deck was written. Fix the problems above and run again${missing === 'error' ? ' (or use --missing keep to leave unfilled tags visible)' : ''}.`);
    return 1;
  }
  printIssues(issues, io, { blankKeys: read.blankKeys });
  io.out(`Wrote ${out}`);
  return 0;
}
